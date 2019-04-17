from gym import Env, spaces
from gym.utils import seeding
from string import Template
import os, sys
import numpy as np
import math
import time
import win32com.client as com
from random import randint
import datetime
from goto import with_goto


class VissimEnv(Env):
    def __init__(self):
        self.Vissim = com.Dispatch("Vissim.Vissim")
        Path = os.getcwd()
        # Load a Vissim Network:
        # Filename = os.path.join(Path, 'test.inpx')
        Filename = r'C:\Users\29904\Desktop\new\test.inpx'
        flag_read_additionally = False  # you can read network(elements) additionally, in this case set "flag_read_additionally" to true
        self.Vissim.LoadNet(Filename, flag_read_additionally)

        # Load a Layout:
        # Filename = os.path.join(Path, 'test.layx')
        Filename = r'C:\Users\29904\Desktop\new\test.layx'
        self.Vissim.LoadLayout(Filename)
        self.speed_limit = 100 / 3.6
        self.sensor_dis = 150
        self.epi = 0
        self.time_step = 0
        self.VissimDebug = self.VissimDebug()
        # Define action sapce and observation space
        self.action_space = spaces.Discrete(int(self.speed_limit) + 1)
        self.observation_space = spaces.Box(low=0, high=1, shape=(14,), dtype=np.float32)

        self.seed()
        self.viewer = None
        self.state = None
        self.input_info = None

        End_of_simulation = 999999  # simulation second [s]
        self.Vissim.Simulation.SetAttValue('SimPeriod', End_of_simulation)

        # Set maximum speed:
        self.Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)

        # Activate QuickMode:
        self.Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode", 1)

        # Simulate so that the new state is active in the Vissim simulation:
        Sim_break_at = 100  # simulation second [s]
        self.Vissim.Simulation.SetAttValue("SimBreakAt", Sim_break_at)

    def seed(self, seed=None):
        self.np_random, seed = seeding.np_random(seed)
        return [seed]

    @with_goto
    def step(self, action):
        assert self.action_space.contains(action), "%r (%s) invalid" % (action, type(action))
        acce, desired_vel, r_t_first = self.acce_output(action)
        # Derive the velocity
        vel = self.input_info["vel"]
        link_coordinate = self.input_info["link_coordinate"]
        if vel + acce * 0.1 < 0:
            vel = 0
        else:
            link_coordinate += vel * 0.1 + 0.5 * acce * 0.1 * 0.1
            vel += acce * 0.1
        # Step 3-4 Execution of CF and LC
        # This function will operate during the next simulation step
        av_link = 1
        av_lane = self.input_info["lane"]
        self.new_Vehicle.MoveToLinkPosition(av_link, av_lane, link_coordinate)
        self.Vissim.Simulation.RunSingleStep()
        # Get the state after executing action
        all_veh_attributes = self.Vissim.Net.Vehicles.GetMultipleAttributes(
            ('No', 'VehType', 'Speed', 'Pos', 'Lane', 'DestLane', 'PosLat'))
        all_veh_attributes = np.asarray(all_veh_attributes)
        av_no = self.new_Vehicle.AttValue('No')
        av_id = np.where(all_veh_attributes[:, 0] == av_no)
        if not len(av_id[0]):
            done = 1
            r_t_first = 0
            observation = []
            reward = 0
            goto.end
        link_coordinate = self.new_Vehicle.AttValue('Pos')
        # Get link and lane num of AV
        av_link_lane = self.new_Vehicle.AttValue('Lane').split('-', 1)
        av_lane = int(av_link_lane[1])
        # Delete the info of AV and any LC car to find nearby vehicles faster
        normal_id = np.where(all_veh_attributes[:, 1] == '100')
        all_veh_attributes = all_veh_attributes[normal_id]
        # Get the neighbor info of AV
        input_info = self.get_info(all_veh_attributes, 1, av_lane, link_coordinate)
        # Get the reward after executing action
        ini_acce = self.input_info["acce_previous"]
        # Dump the info of AV itself
        input_info["vel"] = vel
        input_info["acce_previous"] = acce
        input_info["link_coordinate"] = link_coordinate
        input_info["lane"] = av_lane
        input_info["lane_num"] = 3
        self.input_info = input_info
        reward = self.get_reward(desired_vel, acce, r_t_first, ini_acce)
        observation = self.make_observaton(input_info)
        print("Episode", self.epi, "step", self.time_step, "AV pos=", link_coordinate, "desired vel=", desired_vel,
              "a=",
              acce, "vel=", vel, "gap_lead=", input_info["gap_lead"], "vel diff=",
              input_info["vel"] - input_info["vel_lead"], "Reward", reward)
        # Whether break out to next episode
        if input_info["link_coordinate"] > 999 or input_info["gap_lead"] < 1:
            done = 1
        else:
            done = 0
        label.end
        self.time_step += 1
        return observation, reward, done, r_t_first

    def acce_output(self, action):
        # directly output of desired vel
        desired_vel = action
        update_flag = 1
        # get the state
        input_info = self.input_info
        # for the desired vel is too small OR too large
        r_t_first = 100
        if desired_vel < 5:
            desired_vel = 5
        if desired_vel > self.speed_limit:
            desired_vel = self.speed_limit
        # derive acceleration using IDM
        s0 = 2
        T = 1.5
        a = 0.73
        b = 1.67
        exponent = 4
        if input_info["vel"] < desired_vel:
            acce_base = a
        else:
            acce_base = b
        a_idm = acce_base * (1 - pow(input_info["vel"] / desired_vel, exponent))

        ## hard constrains for acceleration
        # urgent stop
        if input_info["vel_lead"] < 1:
            if input_info["gap_lead"] > s0:
                stop_dis = input_info["gap_lead"] - s0
            else:
                stop_dis = input_info["gap_lead"]
            a_idm = - pow(input_info["vel"], 2) / 2 / stop_dis
            r_t_first = 0
        # no lead car
        if input_info["gap_lead"] > 145:
            desired_vel = self.speed_limit
            a_idm = a * (1 - pow(input_info["vel"] / desired_vel, exponent))
            r_t_first = 0

        # dynamic constraints
        if a_idm < -3:
            a_idm = -3
        if a_idm > 3:
            a_idm = 3
        return a_idm, desired_vel, r_t_first

    def get_reward(self, desired_vel, a_idm, r_t_first, acce_pre):
        input_info = self.input_info
        # dangerous gap
        if input_info["gap_lead"] < 1 * input_info["vel"] and r_t_first != 0:
            r_t_first = -1
        # uncomfortable jerk
        if abs(a_idm - acce_pre) / 0.1 > 3.5:
            r_t_first = -1
        if r_t_first != 100:
            reward = r_t_first
        else:
            reward = input_info["vel"] / self.speed_limit - input_info["gap_lead"] / self.sensor_dis - abs(
                a_idm - acce_pre) / 0.1 / 24
            print('part1=', input_info["vel"] / self.speed_limit, ' part2=', - input_info["gap_lead"] / self.sensor_dis,
                  ' part3=', - abs(a_idm - acce_pre) / 0.1 / 24)
        return reward

    def reset(self):
        if not self.epi:
            print("Vissim Experiment Start.")
            # Choose Random Seed
            Random_Seed = randint(1, 100)
            self.Vissim.Simulation.SetAttValue('RandSeed', Random_Seed)
        else:
            self.Vissim.Simulation.Stop()
        self.epi += 1
        self.time_step = 0
        print("Episode : " + str(self.epi))
        # Step 1 To create a different traffic flow for every episode
        observation = []
        # Set vehicle input:
        VI_number = 1  # VI = Vehicle Input
        # Make ascending volume for 3 lanes：more jam and less free
        if self.epi % 10 == 0:
            new_volume = randint(600, 2160)
        elif self.epi % 10 <= 2:
            new_volume = randint(2400, 3840)
        else:
            new_volume = randint(3900, 5250)
        # else:
        #     new_volume = randint(5400, 6300)
        # vehicles per hour
        self.Vissim.Net.VehicleInputs.ItemByKey(VI_number).SetAttValue('Volume(1)', new_volume)

        # Set vehicle composition:
        Veh_composition_number = 1
        Rel_Flows = self.Vissim.Net.VehicleCompositions.ItemByKey(Veh_composition_number).VehCompRelFlows.GetAll()
        Rel_Flows[0].SetAttValue('VehType', 100)  # Changing the vehicle type
        Rel_Flows[0].SetAttValue('DesSpeedDistr', 60)  # Changing the desired speed distribution
        Rel_Flows[0].SetAttValue('RelFlow', 1)  # Changing the relative flow
        # Rel_Flows[1].SetAttValue('RelFlow', 0.1)  # Changing the relative flow of the 2nd Relative Flow.

        self.Vissim.Simulation.RunContinuous()  # start the simulation until SimBreakAt (100s)
        # Step 2 To create an autonomous vehicle naturally
        # Accessing all attributes directly using "GetMultipleAttributes"
        all_veh_attributes = self.Vissim.Net.Vehicles.GetMultipleAttributes(
            ('No', 'VehType', 'Speed', 'Pos', 'Lane', 'DestLane', 'PosLat', 'Acceleration'))
        all_veh_attributes = np.asarray(all_veh_attributes)
        if not len(all_veh_attributes):
            print('Empty Road!')
            self.Vissim.Simulation.Stop()
        else:
            all_pos = all_veh_attributes[:, 3]
            all_pos = all_pos.astype(np.float)
            # Get the id of one vehicle
            veh_id = np.where(all_pos == min(all_pos))
            veh_id = veh_id[0][0]
            veh_number = int(all_veh_attributes[veh_id][0])
            vel = float(all_veh_attributes[veh_id][2]) / 3.6
            # Remove the vehicle that will be replaced by our AV
            self.Vissim.Net.Vehicles.RemoveVehicle(veh_number)
            # Put our AV to the network
            vehicle_type = 666
            desired_speed = 0  # unit according to the user setting in Vissim [km/h or mph]
            link = 1
            lane = all_veh_attributes[veh_id][4][2]
            xcoordinate = min(all_pos)  # unit according to the user setting in Vissim [m or ft]
            ini_acce = all_veh_attributes[veh_id][7]
            interaction = True  # optional boolean
            new_Vehicle = self.VissimDebug.AddVehicle(self.Vissim, vehicle_type, link, lane, xcoordinate, desired_speed,
                                                      interaction)

            # Step 3-1 Get the necessary info
            all_veh_attributes = self.Vissim.Net.Vehicles.GetMultipleAttributes(
                ('No', 'VehType', 'Speed', 'Pos', 'Lane', 'DestLane', 'PosLat'))
            all_veh_attributes = np.asarray(all_veh_attributes)
            av_no = new_Vehicle.AttValue('No')
            av_id = np.where(all_veh_attributes[:, 0] == av_no)
            if not len(av_id[0]):
                self.Vissim.Simulation.Stop()
            link_coordinate = new_Vehicle.AttValue('Pos')
            # Get link and lane num of AV
            av_link_lane = new_Vehicle.AttValue('Lane').split('-', 1)
            av_lane = int(av_link_lane[1])
            # Delete the info of AV and any LC car to find nearby vehicles faster
            normal_id = np.where(all_veh_attributes[:, 1] == '100')
            all_veh_attributes = all_veh_attributes[normal_id]
            # Get the neighbor info of AV
            input_info = self.get_info(all_veh_attributes, 1, av_lane, link_coordinate)
            # Init the vel of AV
            vel = min(vel, input_info["vel_lead"])
            # Dump the info of AV itself
            input_info["vel"] = vel
            input_info["lane"] = av_lane
            input_info["lane_num"] = 3
            input_info["acce_previous"] = ini_acce
            input_info["link_coordinate"] = link_coordinate
            self.input_info = input_info
            self.new_Vehicle = new_Vehicle
            observation = self.make_observaton(input_info)
        return observation

    def render(self, mode='human', close=False):
        pass

    def close(self):
        self.Vissim.Simulation.Stop()

    def get_info(self, all_veh_attributes, av_link, av_lane, link_coordinate):
        input_info = {}
        all_vel = all_veh_attributes[:, 2]
        all_vel = all_vel.astype(np.float)
        all_pos = all_veh_attributes[:, 3]
        all_pos = all_pos.astype(np.float)
        all_lane = all_veh_attributes[:, 4]
        all_DestLane = all_veh_attributes[:, 5]
        all_PosLat = all_veh_attributes[:, 6]

        # Left side
        all_left_id = np.where(all_lane == str(av_link) + '-' + str(av_lane + 1))
        all_left_pos = all_pos[all_left_id]
        all_left_vel = all_vel[all_left_id]
        if all_left_pos.size == 0:
            pos_leftlead = link_coordinate + self.sensor_dis
            vel_leftlead = self.speed_limit
            lat_leftlead = 1
            pos_leftlag = link_coordinate - self.sensor_dis
            vel_leftlag = 0
            lat_leftlag = 1
            ave_vel_left = self.speed_limit
        else:
            ## Average speed of left lane
            all_sensor_left = np.where(
                (all_left_pos < link_coordinate + self.sensor_dis) & (all_left_pos > link_coordinate - self.sensor_dis))
            all_sensor_left_vel = all_left_vel[all_sensor_left]
            if all_sensor_left_vel.size == 0:
                ave_vel_left = self.speed_limit
            else:
                ave_vel_left = np.mean(all_sensor_left_vel) / 3.6
            # Left lead
            all_left_pos = all_left_pos.tolist()
            all_left_pos.sort()
            pos_leftlead = next((x for x in all_left_pos if x > link_coordinate), None)
            if not pos_leftlead:
                pos_leftlead = link_coordinate + self.sensor_dis
                vel_leftlead = self.speed_limit
                lat_leftlead = 1
            elif float(pos_leftlead) < link_coordinate + self.sensor_dis:
                pos_leftlead = float(pos_leftlead)
                leftlead_id = np.where(all_pos == pos_leftlead)
                leftlead_id = leftlead_id[0][0]
                vel_leftlead = all_vel[leftlead_id] / 3.6
                lat_leftlead = self.latShift(all_DestLane[leftlead_id], av_lane + 1, av_lane,
                                             all_PosLat[leftlead_id])
            else:
                pos_leftlead = link_coordinate + self.sensor_dis
                vel_leftlead = self.speed_limit
                lat_leftlead = 1
            # Left lag
            all_left_pos.sort(reverse=True)
            pos_leftlag = next((x for x in all_left_pos if x < link_coordinate), None)
            if not pos_leftlag:
                pos_leftlag = link_coordinate - self.sensor_dis
                vel_leftlag = 0
                lat_leftlag = 1
            elif float(pos_leftlag) > link_coordinate - self.sensor_dis:
                pos_leftlag = float(pos_leftlag)
                leftlag_id = np.where(all_pos == pos_leftlag)
                leftlag_id = leftlag_id[0][0]
                vel_leftlag = all_vel[leftlag_id] / 3.6
                lat_leftlag = self.latShift(all_DestLane[leftlag_id], av_lane + 1, av_lane,
                                            all_PosLat[leftlag_id])
            else:
                pos_leftlag = link_coordinate - self.sensor_dis
                vel_leftlag = 0
                lat_leftlag = 1

        # Right side
        all_right_id = np.where(all_lane == str(av_link) + '-' + str(av_lane - 1))
        all_right_pos = all_pos[all_right_id]
        all_right_vel = all_vel[all_right_id]
        if all_right_pos.size == 0:
            pos_rightlead = link_coordinate + self.sensor_dis
            vel_rightlead = self.speed_limit
            lat_rightlead = 1
            pos_rightlag = link_coordinate - self.sensor_dis
            vel_rightlag = 0
            lat_rightlag = 1
            ave_vel_right = self.speed_limit
        else:
            ## Average speed of right lane
            all_sensor_right = np.where(
                (all_right_pos < link_coordinate + self.sensor_dis) & (
                        all_right_pos > link_coordinate - self.sensor_dis))
            all_sensor_right_vel = all_right_vel[all_sensor_right]
            if all_sensor_right_vel.size == 0:
                ave_vel_right = self.speed_limit
            else:
                ave_vel_right = np.mean(all_sensor_right_vel) / 3.6
            # Right lead
            all_right_pos = all_right_pos.tolist()
            all_right_pos.sort()
            pos_rightlead = next((x for x in all_right_pos if x > link_coordinate), None)
            if not pos_rightlead:
                pos_rightlead = link_coordinate + self.sensor_dis
                vel_rightlead = self.speed_limit
                lat_rightlead = 1
            elif float(pos_rightlead) < link_coordinate + self.sensor_dis:
                pos_rightlead = float(pos_rightlead)
                rightlead_id = np.where(all_pos == pos_rightlead)
                rightlead_id = rightlead_id[0][0]
                vel_rightlead = all_vel[rightlead_id] / 3.6
                lat_rightlead = self.latShift(all_DestLane[rightlead_id], av_lane - 1, av_lane,
                                              all_PosLat[rightlead_id])
            else:
                pos_rightlead = link_coordinate + self.sensor_dis
                vel_rightlead = self.speed_limit
                lat_rightlead = 1
            # Right lag
            all_right_pos.sort(reverse=True)
            pos_rightlag = next((x for x in all_right_pos if x < link_coordinate), None)
            if not pos_rightlag:
                pos_rightlag = link_coordinate - self.sensor_dis
                vel_rightlag = 0
                lat_rightlag = 1
            elif float(pos_rightlag) > link_coordinate - self.sensor_dis:
                pos_rightlag = float(pos_rightlag)
                rightlag_id = np.where(all_pos == pos_rightlag)
                rightlag_id = rightlag_id[0][0]
                vel_rightlag = all_vel[rightlag_id] / 3.6
                lat_rightlag = self.latShift(all_DestLane[rightlag_id], av_lane - 1, av_lane,
                                             all_PosLat[rightlag_id])
            else:
                pos_rightlag = link_coordinate - self.sensor_dis
                vel_rightlag = 0
                lat_rightlag = 1

        # The lane of AV
        all_av_id = np.where(all_lane == str(av_link) + '-' + str(av_lane))
        all_av_pos = all_pos[all_av_id].tolist()
        # Lead
        all_av_pos.sort()
        pos_lead = next((x for x in all_av_pos if x > link_coordinate), None)
        if not pos_lead:
            pos_lead = link_coordinate + self.sensor_dis
            vel_lead = self.speed_limit
        elif pos_lead < link_coordinate + self.sensor_dis:
            pos_lead = float(pos_lead)
            lead_id = np.where(all_pos == pos_lead)
            lead_id = lead_id[0][0]
            vel_lead = all_vel[lead_id] / 3.6
        else:
            pos_lead = link_coordinate + self.sensor_dis
            vel_lead = self.speed_limit
        # Lag
        all_av_pos.sort(reverse=True)
        pos_lag = next((x for x in all_av_pos if x < link_coordinate), None)
        if not pos_lag:
            pos_lag = link_coordinate - self.sensor_dis
            vel_lag = 0
        elif pos_lag > link_coordinate - self.sensor_dis:
            pos_lag = float(pos_lag)
            lag_id = np.where(all_pos == pos_lag)
            lag_id = lag_id[0][0]
            vel_lag = all_vel[lag_id] / 3.6
        else:
            pos_lag = link_coordinate - self.sensor_dis
            vel_lag = 0
        # Dump all the input info
        input_info["gap_leftlead"] = float(pos_leftlead) - link_coordinate - 4.25
        input_info["gap_leftlag"] = link_coordinate - float(pos_leftlag) - 4.25
        input_info["gap_rightlead"] = float(pos_rightlead) - link_coordinate - 4.25
        input_info["gap_rightlag"] = link_coordinate - float(pos_rightlag) - 4.25
        input_info["gap_lead"] = float(pos_lead) - link_coordinate - 4.25
        input_info["gap_lag"] = link_coordinate - float(pos_lag) - 4.25
        input_info["lat_leftlead"] = lat_leftlead
        input_info["lat_leftlag"] = lat_leftlag
        input_info["lat_rightlead"] = lat_rightlead
        input_info["lat_rightlag"] = lat_rightlag
        input_info["vel_rightlead"] = vel_rightlead
        input_info["vel_rightlag"] = vel_rightlag
        input_info["vel_leftlead"] = vel_leftlead
        input_info["vel_leftlag"] = vel_leftlag
        input_info["vel_lead"] = vel_lead
        input_info["vel_lag"] = vel_lag
        # Add the average vel of left and right side
        input_info["ave_vel_left"] = ave_vel_left
        input_info["ave_vel_right"] = ave_vel_right
        return input_info

    def latShift(self, Destlane, lc_lane, av_lane, Poslat):
        if Destlane == av_lane:
            if lc_lane != Destlane:
                latshift = 1 - abs(Poslat - 0.5)
            else:
                latshift = abs(Poslat - 0.5)
        else:
            latshift = 1
        # If the same lane as AV
        if lc_lane == av_lane:
            latshift = abs(Poslat - 0.5)
        return latshift

    def make_observaton(self, input_info):
        gap_leftlead = input_info["gap_leftlead"]
        gap_leftlag = input_info["gap_leftlag"]
        gap_rightlead = input_info["gap_rightlead"]
        gap_rightlag = input_info["gap_rightlag"]
        gap_lead = input_info["gap_lead"]
        gap_lag = input_info["gap_lag"]
        vel_rightlead = input_info["vel_rightlead"]
        vel_rightlag = input_info["vel_rightlag"]
        vel_leftlead = input_info["vel_leftlead"]
        vel_leftlag = input_info["vel_leftlag"]
        vel_lead = input_info["vel_lead"]
        vel_lag = input_info["vel_lag"]
        vel = input_info["vel"]
        lane = input_info["lane"]
        lane_num = input_info["lane_num"]
        standard_dis = self.sensor_dis
        standard_vel = self.speed_limit
        observation = np.array([gap_leftlead / standard_dis, gap_leftlag / standard_dis, gap_rightlead / standard_dis,
                                gap_rightlag / standard_dis, gap_lead / standard_dis, gap_lag / standard_dis,
                                vel_rightlead / standard_vel, vel_rightlag / standard_vel, vel_leftlead / standard_vel,
                                vel_leftlag / standard_vel, vel_lead / standard_vel, vel_lag / standard_vel,
                                vel / standard_vel, lane / lane_num])
        return observation

    class VissimDebug():
        def AddVehicle(self, Vissim, vehicle_type, link, lane, xcoordinate, desired_speed, interaction):
            try:
                new_Vehicle = Vissim.Net.Vehicles.AddVehicleAtLinkPosition(vehicle_type, link, lane, xcoordinate,
                                                                           desired_speed, interaction)
            except:
                print("COM ERROR")
            return new_Vehicle

        def MoveVehicle(self, Vehicle, av_link, lane_number, link_coordinate):
            try:
                Vehicle.MoveToLinkPosition(av_link, lane_number, link_coordinate)
            except:
                print("COM ERROR")
            return Vehicle
